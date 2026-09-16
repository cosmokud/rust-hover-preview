# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.70+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

Rust Hover Preview is a Windows 11 system tray app that shows instant image, video, PDF, and text previews in File Explorer when you hover files with the mouse or navigate with the keyboard.

Inspired by QTTabBar (QuizoApps) hover preview.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer
- Static image previews plus animated GIF, APNG, and libwebp-backed animated WebP playback
- Animation memory is bounded by a sliding window rather than the length of the file: the decoder stays a few frames ahead of the playhead and frames already shown are released, so long or large GIF, APNG, and WebP files play to the end and loop instead of stopping early
- Video previews through FFmpeg (`ffplay` + `ffprobe`)
- PDF previews of the first page, rendered by the PDF engine already built into Windows — no bundled renderer, no extra install, and no new dependency
- Text and code previews colored by syntax definition, with Atom One Light and One Dark Pro bundled as the light and dark themes — or any `.tmTheme` file dropped into the app's own `theme` folder — and Markdown shown as the document it describes (or as highlighted source, from the tray)
- Tray controls for enable/disable, delay, positioning, scaling, startup, off-trigger key, volume, text previews, text theme, text size, and Markdown rendering
- Explorer Shell view detection, folder caching, and path normalization for reliable hover matching
- Topmost, non-activating preview windows designed to avoid focus stealing
- Per-monitor DPI awareness to reduce scaling artifacts on high-DPI displays
- Display-aware placement that keeps the preview inside the monitor under the cursor or focused item, with frames painted before the window is revealed so a new hover never flashes the previous image
- Scale-aware sizing from 25% to 400% or fit-to-screen, always clamped to the space available on the display so a scaled-up preview is never clipped (PDF and text previews have their own sizing, described below)
- Sleep/resume resilience: waking the system resets the layered composition surface and re-asserts the preview's topmost/layered styles, the tray icon is restored after an Explorer restart, and video playback and background decoding are torn down cleanly on suspend
- EnumWindows-based Explorer detection with CabinetWClass/ExplorerWClass class matching to keep idle polling light and avoid Explorer-side COM allocations, plus input-grace helpers that throttle hover and keyboard focus probes to recent user activity
- Wheel scrolling refreshes the preview without moving the mouse: a system-wide low-level mouse hook reports wheel input, the hover stability window restarts while the wheel turns, and the item that settles under the parked cursor is previewed — or the preview is dropped when that item is not media. Scrolling while a keyboard preview is on screen closes it and hands the screen back to the mouse, so the preview follows the item under the cursor instead of staying frozen while the list scrolls
- Keyboard previews take priority over the pointer: a preview opened with the arrow keys stays on top when a parked cursor sits under it, mouse-driven triggers stay frozen until the mouse is moved on purpose or the wheel is turned, and the cursor-over-preview check is skipped instead of running on every poll
- Single-instance enforcement: launching the app again exits immediately instead of adding a duplicate tray icon, Explorer hook, and preview window

## Supported Formats

### Images

`jpg`, `jpeg`, `jpe`, `jfif`, `png`, `apng`, `gif`, `bmp`, `ico`, `tiff`, `tif`, `webp`, `tga`, `pbm`, `pgm`, `ppm`, `pam`, `pnm`, `hdr`, `exr`, `qoi`, `ff`

`apng` files with multiple frames play as animations; single-frame ones show as static images. Animation is detected from the file's content rather than the extension, so a `.png` file carrying an `acTL` chunk ahead of its image data animates as well, and every other `.png` stays a static image.

### PDF

`pdf`

The first page is rendered by the PDF engine that ships with Windows, so there is no renderer to bundle and nothing to install. The page is rendered at the exact size the preview asks for, which keeps a scaled-up preview sharp instead of enlarging a thumbnail, and the page size is read from the file itself, so anything that is not A4 or Letter is placed correctly.

Because the page is drawn at the preview's size, a PDF always takes the largest size the display area allows — `Preview Scaling` does not apply to it, since a bigger preview is sharper text rather than an enlarged image.

Any PDF the Windows engine can open without a password is previewed. Password-protected and damaged files are skipped rather than shown as an error, and a `.pdf` file has to contain a PDF header before it is handed to the renderer.

### Text and code

`txt`, `md`, `rtf`, `nfo`, `json`, `toml`, `yaml`, `xml`, `ini`, `csv`, `log`, `sql`, `py`, `js`, `ts`, `rs`, `go`, `c`, `h`, `cpp`, `cs`, `java`, `kt`, `swift`, `php`, `rb`, `lua`, `sh`, `ps1`, `bat`, `html`, `css`, and the rest of the list below

The first screenful is shown, colored by the syntax definition its extension names — the TextMate grammars `bat` uses, so the colors match what an editor would show. Anything no grammar claims is still previewed, as plain text.

The full list lives in `config.ini` and is meant to be edited there: add an extension the app does not know, or delete the ones you never want to see, and the change applies without a restart. A file whose extension is not in the list is not previewed at all.

The files a repository is recognized by usually have no extension to match: `LICENSE`, `Makefile` and `Dockerfile` have no dot in them anywhere, and `.gitignore` has no extension at all — its dot begins its *name*. So there is a second list beside the extensions, of names, and a file is previewed when either list claims it. Dot files are matched by the name they are written with, with the dot dropped on both sides: `gitignore` in the extension list means `.gitignore` exactly as `.gitignore` in the name list does, and `.eslintrc.json` is read by its extension because that is the part of it the filesystem calls one. The default list covers what a GitHub repository puts at its root — licences and notices, `README` and `CHANGELOG`, the make and container files, and the dot files (`.gitignore`, `.gitattributes`, `.editorconfig`, `.env` and their kind).

`.md` files are rendered as the document they describe — headings, lists, quotes, tables, links, and fenced code blocks colored with the same highlighter — and can be switched to highlighted source from the tray or with `markdown_mode`.

`.rtf` files are shown as the text they carry: Word's tables are skipped and the paragraphs, tabs, and escaped characters are resolved, so a document reads without its formatting.

`.nfo` files are read as CP437 art with their ANSI colors preserved, toned to the page they are drawn on so light and dark themes are both readable.

Text previews are sized to their content rather than to the image scaling setting: a two-line file gets a two-line preview, a long file takes as much of the display as the space beside the cursor allows, and the font is a fixed size scaled by the display's DPI rather than a stretched image. `Text Preview Font Size` sets that size — the glyphs, the line spacing and the margin scale together, so 200% is the same page twice the size and shows fewer lines, not the same lines stretched. Code, markup, and NFO art keep their columns and are clipped at the right edge; prose — a readme, a log, an `.rtf` — wraps.

### Themes

Text and code are colored by one of two bundled TextMate themes — Atom One Light (the default) and One Dark Pro — or by any TextMate theme you drop into `%APPDATA%\rust-hover-preview\theme`, a folder the app creates on first run. A `.tmTheme` file is listed in the `Text Preview Theme` submenu under its own name without the extension, below the two bundled ones, so `atom-one-light.tmTheme` appears as `atom-one-light` — and a file named after a bundled theme is a theme of its own rather than a replacement for it. The folder is read each time the menu is opened, which is when a file added or edited since the last look takes effect. Nothing in the folder is read until a theme is used, so a shelf of files costs nothing while it sits there, and a file that is missing, unreadable, or not a TextMate theme at all is painted with the default rather than leaving the preview uncolored: its entry keeps its place and its name stays in `config.ini`, so repairing the file brings the theme back the next time the menu is opened.

### Scrolling a text preview

When a file is longer than the preview, the preview grows a scrollbar instead of stopping at the bottom of the page. Moving the pointer onto the preview keeps it there — the usual rule is that touching a preview dismisses it — and from there the wheel scrolls the text, or the thumb can be dragged. The pointer can stray around the preview without the preview closing, and the amount it can stray is not the same on every side: reaching a preview is a movement towards it, so the hand is still moving when it arrives, and the edge it arrives at — the one the scrollbar sits at — keeps room past it, forty logical pixels at the display's DPI by default, so a hand that overshoots the thin scrollbar does not take the preview down with it. The edge it came from gets no such room, since the pointer has already stopped there. How much room the far edge gets is the `text_scroll_far_edge_grace_pixels` setting. The preview closes when the pointer leaves that region, or when another file is hovered. Scrolling stops at the end of the document: the last frame is pulled back until it is full, its last line is on screen, and turning the wheel further — or dragging the thumb to the bottom of its track — stays on that frame rather than scrolling the text up out of a mostly empty one. A rendered `.md` document counts its lines the way it draws them, so a paragraph is one line and the heading above it is another; the range it scrolls through is those lines, not the blank lines and wrapped paragraphs of the file behind them.

Selecting works the way it does anywhere else: drag across the text to select it, `Ctrl+C` to copy it, or right-click the preview for **Copy**. What is copied is what is on screen, so a line clipped at the right edge copies the part that was visible, and a Copy with nothing selected takes the whole frame. Dragging the scrollbar clears the selection, since the lines it was measured against have moved.

All of this belongs to **full mode** (`Enable Text Preview Full Mode` in the tray, **off by default**). With it off, a text preview is only something to look at: a frame shows what fits and says how many lines it left out on the last one, and the pointer over it closes it the way it closes any other preview.

Scrolling does not re-read the file. For source files only the lines coming into view are highlighted — the parser's state is carried along as you go, with a checkpoint every few dozen lines so a jump in either direction is bounded — and the text itself is read once, up to the read cap. A rendered Markdown document is walked in one pass (its line breaks are only known once the block before it has been read) but only the visible lines are kept. Scrolling is fastest where it matters most: a hundred-thousand-line source file opens as quickly as a short one.

Text previews can be turned off entirely with `Enable Text Preview`, which leaves the extension list as it is; the list is the other half of the gate, and `false` in `text_preview_enabled` is the same switch from `config.ini`.

Encoding is decided from the file: UTF-8 or UTF-16 byte order marks first, then UTF-8, then the Windows code page (CP437 for NFO art). A file that is not text — an executable or archive with a text extension — is skipped rather than shown as garbage. One preview reads up to 2 MB, and a preview can be scrolled through its first 2000 lines; a file cut off at either bound says so on its last line.

### Videos (FFmpeg required)

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`, `ts`, `m2ts`, `mts`, `mpg`, `mpeg`, `vob`, `3gp`, `ogv`, `rmvb`, `asf`, `divx`, `f4v`, `mxf`, `dv`

All video files supported by FFmpeg work as well — any container it can demux and any raw video stream it can decode (`.tp`, `.tr`, `.tod`, `.wtv`, `.ty`, `.vro`, `.nut`, `.gxf`, `.nsv`, `.ivf`, `.y4m`, `.obu`, `.h264`, `.h265`, `.h266`, `.vc1`, `.av1`, `.avs2`, `.avs3`, and camcorder or game formats such as `.bik`, `.bk2`, `.smk`, `.mve`, `.cpk`, `.thp`, `.usm`, `.pmp`, `.kux`, `.dav`). The complete list lives in `src/video_formats.rs`.

`.ts` and `.mts` are shared with TypeScript sources, so those two are accepted only when the file actually contains MPEG-TS packets; a TypeScript file is not previewed.

## Installation (Recommended)

Each release provides two asset options:

- **`rust-hover-preview_<version>_x64-setup.exe`** — the NSIS installer. Run it to install to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- **`rust-hover-preview.exe`** — the standalone portable binary. Place it in any folder on your PC (for example: `C:\Tools\RustHoverPreview`) and run it directly. No installation needed.

1. Open [Releases](../../releases)
2. Download your preferred asset
3. Run the installer or place the portable binary wherever you like
4. Launch Rust Hover Preview

No Rust toolchain is needed when installing from Releases.

> **Note for existing users upgrading from an earlier version:**  
> The installer handles upgrades automatically, cleaning up the old `%LOCALAPPDATA%\Rust Hover Preview` folder if present.

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
- **Preview Delay**: `Instant (0 ms)`, `Fast (200 ms)`, `Medium (500 ms)`, `Relaxed (750 ms)`, `Slow (1000 ms)`
- **Same File Rehover Delay**: `Instant (0 ms)`, `Fast (200 ms)`, `Medium (500 ms)`, `Relaxed (750 ms)`, `Slow (1000 ms)` — delay before the same file can preview again after the preview self-dismisses
- **Video Volume**: `Max (100%)`, `High (80%)`, `Medium (50%)`, `Low (25%)`, `Very Low (10%)`, `Mute (0%)`
- **Preview Position**: `Follow Cursor` or `Best Position` — `Best Position` places the preview beside the cursor (or the focused item) and centers it on the same line, moving it only as far as a display edge requires, so a small preview appears where you are looking rather than in the middle of the screen. `Follow Cursor` places it in the roomiest quadrant around the cursor.
- **Enable Text Preview**: Turn text and code previews on or off, ahead of the extension list — turning them off leaves the list alone and turning them back on restores it.
- **Enable Text Preview Full Mode**: Whether a text preview can be worked with — scrolled, selected from and copied, and rested on by the pointer — or is only something to look at. Off by default.
- **Text Preview Theme**: `Atom One Light (Default)`, `One Dark Pro`, and every `.tmTheme` file in `%APPDATA%\rust-hover-preview\theme`, listed under its own file name (`atom-one-light`, `one-dark-pro`) below the two bundled themes. The folder is read every time the menu opens, so a file added or edited while the app runs is listed and used without a restart. A file that cannot be read or is not a TextMate theme is painted with the default theme instead; its item stays in the list, so repairing the file brings it back. Changing the theme re-renders the preview that is on screen.
- **Text Preview Font Size**: `100%`, `125% (Default)`, `150%`, `175%`, `200%`, `250%`, `300%`, `400%` — the size the text is drawn at, for rendered Markdown as much as for source files. The whole page scales with it (glyphs, line spacing and margin together), so a bigger preview holds fewer lines rather than the same lines stretched. Any hand-edited value in `config.ini` is honored, including one between these steps.
- **Markdown Preview**: `Rendered (Default)` shows a `.md` file as the document it describes, `Highlighted Source` shows the markup itself with Markdown syntax highlighting.
- **Preview Scaling**: `Fit to Screen`, `400%`, `300%`, `200%`, `150%`, `100% (Default)`, `50%`, `25%` — sizes the preview relative to the image or video's native resolution instead of always showing it verbatim. `Fit to Screen` enlarges the preview as much as the display allows. Any scale that would extend past the screen edge is reduced to fit, so the preview is never clipped, in both `Follow Cursor` and `Best Position` modes. PDF and text previews size themselves (a PDF page is always fit to screen, and text is always drawn at its own font size), so this setting does not apply to them.
- **Transparent Background**: `Transparent`, `Black`, `White`, or `Checkerboard`
- **Trigger Key (Alt)**: The key the trigger watches, by name, and what holding it does — one of **Trigger Key to Disable Preview** (the default: hold the key and nothing previews while it is held) or **Trigger Key to Enable Preview** (the reverse: previews wait for the key, so hold it and they appear as you hover, let go and they stop). Only one of the two is active at a time. The key itself is the `trigger_key` setting, so `alt`, `ctrl`, `shift` and `win` are the usual choices.
- **Confirm File Type**: When enabled, validates file content signatures (magic bytes) against the extension to avoid loading mislabeled files. If previews don't appear for certain files that should be supported, try enabling this option — the app will attempt to decode them by their true content type rather than relying solely on the file extension.
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
same_file_rehover_delay_ms=750
preview_enabled=true
trigger_key=alt
trigger_key_mode=disable
confirm_file_type=false
follow_cursor=false
transparent_background=black
webp_playback_fps=90
video_volume=0
preview_scale=100
theme=light
markdown_mode=rendered
text_preview_enabled=true
text_preview_full_mode=false
text_font_scale=125
text_scroll_far_edge_grace_pixels=40

[text]
extensions=txt,text,log,nfo,md,markdown,json,toml,yaml,py,js,ts,rs,...
names=license,notice,makefile,dockerfile,gitignore,.gitattributes,...
```

- `theme` is the color theme for text and code previews: `light` (Atom One Light, the default) or `dark` (One Dark Pro). `atom one light` and `one dark pro` are accepted as well. A theme from the `theme` folder is written by the tray as `custom:<name>`, the marker being what keeps a file named `light.tmTheme` apart from the bundled `light` and the other way round; a value you write yourself is read on the same terms — a bundled name means the bundled theme, and any other name means the file of that name in the `theme` folder (extension optional, case not significant), with a name that is neither leaving the theme as it was.
- `markdown_mode` is how `.md` files are drawn: `rendered` (the default document view) or `source` (the markup with syntax highlighting).
- `text_preview_enabled` is the `Enable Text Preview` toggle: `false` stops text previews without touching the extension list.
- `text_preview_full_mode` is the `Enable Text Preview Full Mode` toggle. It is off unless it is turned on, because it changes what a preview does rather than what it shows: `true` lets a text preview be scrolled, selected from and copied.
- `text_font_scale` is the size text is drawn at, as a percentage between 1 and 1000 (`150` or `150%`; `0` resets to the default of 125). The tray offers steps from 100% to 400%, and a value between those steps is used as written.
- `text_scroll_far_edge_grace_pixels` is how far past the far edge of a text preview in full mode the pointer still counts as using it: a distance in logical pixels between 0 and 1000, 40 by default, scaled by the display's DPI so it is the same distance under a hand at any text size (decimals are honored, and a plain `0` ends the region at the preview itself). It is the room a hand that overshoots the scrollbar as it arrives needs; only the edge the pointer travelled towards gets it.
- `extensions` is every extension previewed as text. It is written in full when the file is created and is shortened in this example. Append an extension to preview one the app does not know, or delete entries to stop previewing them — the change is picked up without a restart. An extension is written without its dot (`py`, not `.py`), and an empty list turns text previews off, so the built-in list comes back only when the key itself is missing. A dot file is named here without its dot too, so `gitignore` covers `.gitignore`.
- `names` is the other half of the text gate: the files that have no extension to match, like `LICENSE`, `Makefile` and `Dockerfile`. Entries are file names, matched whole and without regard to case; the list covers the root of a typical repository, and `gitignore` and `.gitignore` are the same entry here as they are there.
- When the trigger key is held, previews either stop or start, according to `trigger_key_mode`: `disable` keeps previews hidden while it is down, and `enable` shows them only while it is down. The key itself is `trigger_key`.
- When `confirm_file_type` is enabled, the app validates file content signatures (magic bytes) against the extension — useful for files with incorrect extensions.
- `webp_playback_fps` controls the maximum playback speed for animated WebP files (1–90 FPS; 0 resets to the default of 90).
- `preview_scale` controls the preview size relative to the media's native resolution. Use `fit` (or `fit to screen`) to scale the preview as large as the display area allows, or a percentage between 1 and 1000 written as `200`, `200%`, or `75`. A scale larger than the available space is reduced to fit so the preview cannot be clipped; `0` resets to the default of 100.

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

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview.

- Uses Windows accessibility APIs (MSAA + UI Automation) to resolve hovered/focused Explorer items
- Uses Shell COM APIs to identify active Explorer windows and folders
- Uses GDI for image rendering in a layered topmost preview window
- Colors text and code with the TextMate grammars `bat` uses (through `syntect` and `two-face`) and lays the styled lines out into a GDI surface with Consolas, the same layered window, and the same frame shape images arrive in
- Renders Markdown with `pulldown-cmark`, with the theme's own colors asked for by scope, so a rendered document and a highlighted source file share one palette
- Bundles Atom One Light and One Dark Pro as TextMate themes converted from their VS Code sources (see `assets/themes/NOTICE.md`), and reads any `.tmTheme` dropped into `%APPDATA%\rust-hover-preview\theme` with the same reader
- Uses Google's libwebp through `webp-animation` for animated WebP decoding
- Uses `directories` for Windows roaming configuration paths
- Uses `ffprobe` for video dimensions and `ffplay` for video playback
- Sets per-monitor DPI awareness (v2 with fallback) on startup to prevent scaling artifacts on layered windows
- Bounds the preview to the monitor under the cursor or focused item via `MonitorFromPoint`/`GetMonitorInfoW` (with a virtual-screen fallback) and repositions the window before installing a frame, so cross-display hops at different scale never strand a stale preview
- Uses the registry (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) for startup control
- Enforces a single running instance with a session-local named mutex, exiting before COM setup, hooks, or background threads start
- Counts and classifies Explorer browser windows via EnumWindows and CabinetWClass/ExplorerWClass class matching, so idle polling never spins up Explorer's shell automation providers
- Gates hover and keyboard focus probes behind input-grace windows (recent_elapsed_within, should_probe_keyboard_focus, should_probe_hover_resolver, should_probe_stationary_hover) and a stationary_hover_probe_done latch to avoid redundant accessibility work for a parked cursor
- Watches wheel input with a low-level mouse hook (`WH_MOUSE_LL`) on its own message-pumping thread, so a scroll that moves the list under a parked pointer re-resolves the hovered item instead of freezing the preview until the mouse moves

## License

MIT. See [LICENSE](LICENSE).

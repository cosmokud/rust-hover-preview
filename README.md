# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.88+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar that shows instant File Explorer previews when you hover a file with the mouse or navigate with the arrow keys, and the preview appears beside it.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Features

- Mouse-hover and keyboard-navigation previews in Explorer
- Images, including animated GIF, APNG, and WebP
- Videos through FFmpeg
- PDF first pages via the built-in Windows PDF engine
- Text and code with syntax highlighting, rendered Markdown, and bundled/custom themes
- Archive contents — zip, rar, 7z, tar — as a file tree with sizes, read without unpacking anything
- Preview scaling from 25% to 400%, or fit-to-screen
- Previews appear beside the cursor or focused item and are never clipped by screen edges
- Tray menu and hand-editable `config.ini`
- DPI aware, single-instance, sleep/resume resilient, and light on idle CPU

## Supported Formats

### Images

`jpg`, `jpeg`, `png`, `apng`, `gif`, `bmp`, `ico`, `tiff`, `webp`, `tga`, `hdr`, `exr`, `qoi`, and more. Animated GIF, APNG, and WebP files play; animation is detected from file content.

### PDF

`pdf` — the first page is rendered by the Windows PDF engine. Password-protected and damaged files are skipped.

### Text and code

`txt`, `md`, `rtf`, `nfo`, `json`, `toml`, `yaml`, `xml`, `ini`, `csv`, `log`, `sql`, `py`, `js`, `ts`, `rs`, `go`, `c`, `h`, `cpp`, `cs`, `java`, `kt`, `swift`, `php`, `rb`, `lua`, `sh`, `ps1`, `bat`, `html`, `css`, and more.

Extensionless repository files such as `LICENSE`, `Makefile`, `Dockerfile`, and `.gitignore` are also supported. Text previews show the first screenful, with long lines wrapped instead of cut off. Markdown can be rendered or shown as source. Full mode adds scrolling, selection, copy; it is off by default.

You can add custom extensions via `config.ini`.

### Archives

`zip`, `zipx`, `jar`, `apk`, `xpi`, `cbz`, `rar`, `7z`, `tar`, `tgz`, and `tar.gz`.

A preview lists what the archive holds — a summary line, then a tree of its folders and files with each file's size, folders first and a `… and N more items` line when the listing is long. Nothing is unpacked: only the archive's own table of contents is read, so a preview of a five-gigabyte archive costs the same as a small one. Encrypted archives are marked, and one whose file table is encrypted says so instead of guessing. Add or remove formats through `config.ini`; the list is editable, so a container this list does not name — a `.docx`, say — can be added as the zip it is.

### Videos (FFmpeg required)

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`, `ts`, `m2ts`, `mts`, `mpg`, `mpeg`, `vob`, `3gp`, `ogv`, `rmvb`, `asf`, `divx`, `f4v`, `mxf`, `dv`. FFmpeg-supported containers and codecs generally work. `.ts` and `.mts` are previewed as video only when they contain MPEG-TS packets.

### Themes

Text, code, and archive listings use Atom One Light (default), One Dark Pro, or any `.tmTheme` file placed in `%APPDATA%\rust-hover-preview\theme`. Archive listings follow the tray's **Text Preview → Font Size** setting.

## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

1. Open [Releases](../../releases)
2. Download your preferred asset
3. Run the installer, or place the portable binary wherever you like
4. Launch Rust Hover Preview

No Rust toolchain is needed. If upgrading from an earlier version, the installer cleans up the old `%LOCALAPPDATA%\Rust Hover Preview` folder automatically.

## Optional: Enable Video Preview (FFmpeg)

Video previews need `ffplay` and `ffprobe` in your `PATH`.

**Option A: winget**

```powershell
winget install --id Gyan.FFmpeg -e
```

**Option B: manual**

1. Download a Windows build from https://ffmpeg.org/download.html
2. Extract it, for example to `C:\ffmpeg`
3. Add `C:\ffmpeg\bin` to your user `PATH`

Verify either way:

```powershell
ffplay -version
ffprobe -version
```

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard (arrow keys/Tab) to preview the focused item.
4. Right-click the tray icon to configure behavior.

## System Tray Menu

- **Enable Preview** — turn previews on or off
- **Preview Types** — Images, Videos, Text, PDF, Archives: switch a kind of preview off without touching its file list
- **Background** — Transparent, Black, White, or Checkerboard
- **Confirm File Type** — validate file content against extension
- **Trigger Key (Alt)** — hold to disable or enable previews
- **Text Preview**
  - **Full Mode** — adds scrolling, selection, and copy; off by default
  - **Theme** — Atom One Light, One Dark Pro, or custom `.tmTheme`
  - **Font Size** — 100%–400%
  - **Markdown** — Rendered or Source
- **Timing**
  - **Delay** — Instant, Fast, Medium, Relaxed, Slow
  - **Rehover Delay** — delay before the same file can preview again
- **Placement**
  - **Position** — Follow Cursor or Best Position, and whether to keep previews off the hovered item's name
  - **Scaling** — Fit to Screen or 25%–400%
- **Volume** — Max, High, Medium, Low, Very Low, Mute
- **Run at Startup** — add or remove the Windows startup entry
- **Config.ini** — open the configuration file; the item is named for the running version
- **Exit** — close the app

## Configuration

Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

The file is watched, and changes apply without a restart.

Example:

```ini
[settings]
run_at_startup=true
hover_delay_ms=0
same_file_rehover_delay_ms=750
preview_enabled=true
image_preview_enabled=true
video_preview_enabled=true
text_preview_enabled=true
pdf_preview_enabled=true
archive_preview_enabled=true
trigger_key=alt
trigger_key_mode=disable
confirm_file_type=false
follow_cursor=false
avoid_filename=true
transparent_background=black
video_volume=0
preview_scale=100
theme=light
markdown_mode=rendered
text_preview_enabled=true
text_preview_full_mode=false
text_font_scale=125

[text]
extensions=txt,text,log,nfo,md,markdown,json,toml,yaml,py,js,ts,rs,...
names=license,notice,makefile,dockerfile,gitignore,.gitattributes,...

[archive]
extensions=zip,zipx,jar,apk,xpi,cbz,rar,7z,tar,tgz,tar.gz
```

Key settings:

- `theme` — `light` (default), `dark`, or a custom theme as `custom:<name>`.
- `markdown_mode` — `rendered` or `source`.
- `image_preview_enabled` / `video_preview_enabled` / `text_preview_enabled` / `pdf_preview_enabled` / `archive_preview_enabled` — whether previews of that kind may be shown at all, without changing the lists of files it covers.
- `text_preview_full_mode` — `true` adds scrolling, selection, and copy.
- `text_font_scale` — percentage from 1 to 1000; default is `125`. Archive listings follow it too.
- `extensions` / `names` — text-preview gates. Extensions are written without dots; names match extensionless files.
- `archive_extensions` — the archive-preview gate, under `[archive]`. Entries are written without dots, and an entry with a dot in it (`tar.gz`) is matched against the end of the file name.
- `trigger_key` / `trigger_key_mode` — key (`alt`, `ctrl`, `shift`, `win`) and mode (`disable` or `enable`).
- `follow_cursor` — `true` for Follow Cursor, `false` for Best Position.
- `avoid_filename` — `true` (the default) keeps a preview off the name of the file it is about, moving it — and, where the display leaves no room beside the name, resizing it — so the item under the pointer or the keyboard stays readable while its preview is up. Applies to both positions.
- `preview_scale` — percentage or `fit`.

## Build from Source

Requirements: Windows 11, Rust 1.88+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

```bash
cargo build            # debug
cargo build --release  # release
```

The release binary is written to `target/release/rust-hover-preview.exe`.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item, GDI paints the preview into a topmost layered window, text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives' own tables of contents, and FFmpeg handles video.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## License

MIT. See [LICENSE](LICENSE).

RAR archives are read with RARLAB's UnRAR sources, compiled into the binary by the `unrar` crate. UnRAR source code may be used in any software to handle RAR archives without limitations and free of charge, but it may not be used to develop a RAR-compatible archiver or to recreate the RAR compression algorithm, which is proprietary. See [RARLAB's licence](https://www.rarlab.com/license.htm) for the full terms.

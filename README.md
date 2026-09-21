# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.88+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar that shows instant File Explorer previews when you hover a file with the mouse or navigate with the arrow keys, and the preview appears beside it.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Features

- Mouse-hover and keyboard-navigation previews in Explorer
- Images, including animated GIF, APNG, and WebP
- HEIC, AVIF, and JPEG XL, decoded by the codec Windows already has
- SVG vectors drawn at the size they are shown, animated or still
- Fonts — `ttf`, `otf`, `ttc`, `woff`, `woff2` — drawn as a specimen: the name the font calls itself and the sample lines its own character map covers
- Videos through FFmpeg, or through Windows' own media engine where FFmpeg is not installed
- PDF first pages via the built-in Windows PDF engine
- Text and code with syntax highlighting, rendered Markdown, and bundled/custom themes
- Archive contents — zip, rar, 7z, tar — as a file tree with sizes, read without unpacking anything
- Office documents — Word, Excel and PowerPoint — drawn from a page Office renders in the background and keeps, so a document previews from its first hover and instantly after
- Preview scaling from 25% to 400%, or fit-to-screen, with a display share of its own for SVG, PDF, Office pages, and fonts
- Previews appear beside the cursor or focused item and are never clipped by screen edges
- Tray menu and hand-editable `config.ini`
- DPI aware, single-instance, sleep/resume resilient, and light on idle CPU

## Supported Formats

You can add or remove formats in config.ini. Unsupported formats will not show a preview except text, which the app will try to force-read.

### Images

`jpg`, `jpeg`, `png`, `apng`, `gif`, `bmp`, `ico`, `tiff`, `webp`, `tga`, `hdr`, `exr`, `qoi`, `heic`, `heif`, `avif`, `jxl`, and more. Animated GIF, APNG, and WebP files play; animation is detected from file content. The last four, and a still `webp`, are decoded by a codec extension Windows provides where one is installed — or, for `webp`, by the libwebp the app carries when it is not; see [Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview](#optional-enable-heic-avif-jpeg-xl-and-webp-preview-windows-codecs).

### Vectors

`svg`, `svgz` — drawn at the size they are shown by the WebView2 runtime Windows 11 ships with, so they stay sharp when enlarged, and animated documents play too. That runtime is what draws a document, so a machine without it shows no SVG preview; nothing of the drawing is done by the app itself.

### Fonts

`ttf`, `otf`, `ttc`, `woff`, `woff2` — drawn by the same WebView2 runtime.

### Videos

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`, `ts`, `m2ts`, `mts`, `mpg`, `mpeg`, `vob`, `3gp`, `ogv`, `rmvb`, `asf`, `divx`, `f4v`, `mxf`, `dv`.

With [FFmpeg installed](#optional-enable-video-preview-ffmpeg) the app plays them through it, so FFmpeg-supported containers and codecs generally work. Without it, videos are played by the media engine Windows already has, which covers `mp4`, `mov`, `m4v`, `mkv`, `webm`, `avi`, `wmv`, `asf`, `ts`, `m2ts`, `mts` and `3gp` — and everything else the codec extensions listed under [Optional: Enable More Video Codecs](#optional-enable-more-video-codecs-windows-codecs) have added. A format neither engine can read shows no preview. The tray's **Codecs** menu says which of the two you have and which codecs are installed.

### PDF

`pdf` — the first page is rendered by the Windows PDF engine. Password-protected and damaged files are skipped.

### Text and code

`txt`, `md`, `rtf`, `nfo`, `json`, `toml`, `yaml`, `xml`, `ini`, `csv`, `log`, `sql`, `py`, `js`, `ts`, `rs`, `go`, `c`, `h`, `cpp`, `cs`, `java`, `kt`, `swift`, `php`, `rb`, `lua`, `sh`, `ps1`, `bat`, `html`, `css`, and more.

Extensionless repository files such as `LICENSE`, `Makefile`, `Dockerfile`, and `.gitignore` are also supported. Markdown can be rendered or shown as source. Full mode adds scrolling, selection, copy;

### Archives

`zip`, `zipx`, `jar`, `apk`, `xpi`, `cbz`, `rar`, `7z`, `tar`, `tgz`, and `tar.gz`.

### Office documents

`doc`, `docm`, `docx`, `dot`, `dotm`, `dotx`, `xls`, `xlsb`, `xlsm`, `xlsx`, `xlt`, `xltm`, `xltx`, `ppt`, `pptm`, `pptx`, `pps`, `ppsm`, `ppsx`, `pot`, `potm`, `potx`.

To show previews for these Office document types, the user must have **Microsoft Office** installed. Excel needs a print queue to export a page — **Microsoft Print to PDF** is enough (make sure that the Print Spooler service is enabled) — and falls back to the sheet's top-left corner without one.

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

Videos preview without FFmpeg, through the media engine Windows ships with — FFmpeg is what makes previews cover far more of them. Install it if you want the formats and codecs Windows does not decode. With it, `ffplay` and `ffprobe` need to be in your `PATH`.

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

## Optional: Enable More Video Codecs (Windows Codecs)

The media engine decodes H.264, MPEG-4 and WMV out of the box, and each of the codecs below is a separate free extension from the Microsoft Store:

| Codec                                  | Needs                                                                                                                      |
| -------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)                                                    |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)                                                     |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)                                                      |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4)                                                   |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)                                                     |

All of them are free, and a Windows 11 device usually has the HEVC, VP9 and AV1 ones already. Installing one takes effect the next time the tray's **Codecs** menu is opened — there is no restart, and nothing to configure. They are not needed at all while FFmpeg is installed.

## Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

`heic`, `heif`, `avif`, `jxl` and a still `webp` are decoded by a codec Windows provides rather than by one shipped with the app, so each one needs its extension installed once from the Microsoft Store:

| Format         | Needs                                                                                                                                            |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |

All of them are free, and a Windows 11 device often has the HEIF, AV1 and WebP ones already. Where one is missing, hovering such a file shows no preview rather than an error, and a multi-image file — a HEIC burst, an animated AVIF, an animated JPEG XL — shows its first frame. A `.webp` is the exception in both directions: it is the one picture here that needs none of them, because the app carries a libwebp decoder of its own — that is what plays an animated one, and what decodes a still one where the WebP codec is missing — so WebP previews on a Windows 10 machine.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard (arrow keys/Tab) to preview the focused item.
4. Right-click the tray icon to configure behavior.

## System Tray Menu

- **Enable Preview** — turn previews on or off.
- **Preview Types** — Images, Videos, Text, PDF, Archives, Office, SVG, Fonts: gate a kind without touching its file list.
- **Text Preview**
  - **Full Mode** — adds scrolling, selection, and copy; off by default.
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in the theme folder.
  - **Font Size** — 400% at the top down to 70% at the bottom.
  - **Markdown** — Rendered or Source.
- **Timing**
  - **Trigger Key (Alt)** — the key is named in the item itself.
    - **Enable Trigger Key** — whether the key is watched at all.
    - **Hold to Disable Preview** / **Hold to Enable Preview** — what holding it does.
  - **Delay** — Instant, Fast, Medium, Relaxed, Slow: 0 ms to 1000 ms.
  - **Rehover Delay** — the same steps, before the same file can preview again.
- **Placement**
  - **Position** — Follow Cursor or Best Position.
  - **Avoid** — Don't Avoid, Avoid Filename (the default), Avoid Filename Column, or Avoid Details: what a preview is kept off instead of covering the item it is about. Avoid Filename keeps it off the hovered name alone, measured as far as the name is actually drawn — extension and all, at the size the view draws it, so a `Content` row's larger name counts too — leaving the rest of the row free to be covered; Avoid Filename Column keeps it off the whole column the name sits in; Avoid Details keeps it off every column a row draws; and Don't Avoid places it by the position alone.
  - **Scaling** — Fit to Screen or 25%–400%.
  - **SVG Scaling** — Fit to Screen, or 75%, 50% (default), 25%, 10% of the display.
  - **PDF Scaling** — Fit to Screen (default), or the same shares of the display.
  - **Office Scaling** — the same for a page Office rendered; a workbook's fallback bitmap is never enlarged.
  - **Font Scaling** — the same shares for a font specimen, 50% (default).
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard.
  - **SVG Background** — the same backdrops for documents.
  - **Font Background** — the same backdrops for a font specimen, which is drawn on a page of its own.
- **Volume** — Max, High, Medium, Low, Very Low, Mute: 100% down to 0%.
- **Performance** — what the app costs to stay fast.
  - **Confirm File Type** — validate file content against the extension.
  - **Office Engine TTL** — how long a family's Office app is kept warm: Indefinitely, 1 hour, 30 minutes, 10 minutes (default), 5 minutes, 1 minute, 0 seconds.
  - **SVG Engine TTL** — how long the browser that draws an SVG document is kept warm; greyed out where WebView2 is missing.
  - **Cache** — memory held between hovers, 2 GB down to 0 MB, each cache's own default marked:
    - **Image** — decoded image frames.
    - **Text** — frames text previews were painted as.
    - **PDF** — pages PDF previews were drawn as.
    - **Office** — pages Office rendered.
  - **Decode Budget** — 16 GB down to 512 MB, 1 GB (default): a file past it gets no preview.
- **Codecs** — what this machine has, for reading only; a check or a cross per row, and the missing ones greyed.
  - **Video** — FFmpeg, Windows Media Foundation, and the video codecs.
  - **Images** — HEIF (HEIC), AVIF, JPEG XL, WebP.
  - **Engines** — WebView2, Microsoft Word, Excel and PowerPoint, and the Windows PDF engine.
- **Run at Startup** — add or remove the Windows startup entry.
- **Config.ini** — open the configuration file; the item is named for the running version.
- **Exit** — close the app.

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
office_preview_enabled=true
svg_preview_enabled=true
font_preview_enabled=true
office_cache_mb=64
office_engine_idle=600
pdf_cache_mb=32
text_cache_mb=0
image_cache_mb=32
decode_budget_gb=1
trigger_key=alt
trigger_key_mode=disable
trigger_key_enabled=true
confirm_file_type=false
follow_cursor=false
avoid_mode=details
image_background=black
svg_background=black
font_background=black
video_volume=0
preview_scale=100
svg_scale=50
pdf_scale=fit
office_scale=fit
font_scale=50
theme=light
markdown_mode=rendered
text_preview_enabled=true
text_preview_full_mode=false
text_font_scale=125

[image]
extensions=apng,avif,bmp,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,svg,svgz,tga,tif,tiff,webp

[video]
extensions=264,265,266,3g2,3gp,3gpp,apv,asf,av1,avc,avi,avs,avs2,avs3,bik,bk2,c93,cavs,cdg,cdxl,cin,cpk,dav,...

[text]
extensions=txt,text,log,nfo,md,markdown,json,toml,yaml,py,js,ts,rs,...
names=license,notice,makefile,dockerfile,gitignore,.gitattributes,...

[archive]
extensions=7z,apk,cbz,jar,rar,tar,tar.gz,tgz,xpi,zip,zipx

[office]
extensions=doc,docm,docx,dot,dotm,dotx,pot,potm,potx,pps,ppsm,ppsx,ppt,pptm,pptx,xls,xlsb,xlsm,xlsx,xlt,xltm,xltx

[font]
extensions=otf,ttc,ttf,woff,woff2
```

Key settings:

- `theme` — `light` (default), `dark`, or `custom:<name>`.
- `markdown_mode` — `rendered` or `source`.
- `image_preview_enabled` / `video_preview_enabled` / `text_preview_enabled` / `pdf_preview_enabled` / `archive_preview_enabled` / `office_preview_enabled` / `svg_preview_enabled` / `font_preview_enabled` — whether previews of that kind may be shown at all; the lists of files it covers stay as they are.
- `text_preview_full_mode` — `true` adds scrolling, selection, and copy.
- `text_font_scale` — a percentage from 1 to 1000, default `125`; archive listings follow it too.
- `extensions` / `names` — the text-preview gates: extensions are written without dots, and a name matches an extensionless file.
- `image_extensions` — the image-preview gate, under `[image]`, written without dots.
- `video_extensions` — the video-preview gate, under `[video]`, written without dots.
- `archive_extensions` — the archive-preview gate, under `[archive]`, written without dots; an entry with a dot in it (`tar.gz`) is matched against the end of the file name.
- `image_cache_mb` — the memory decoded image frames are kept in: default `32`, at most `2048`; a hit skips a full-resolution decode, and `0` holds nothing.
- `office_cache_mb` — the same for the pages Office rendered: default `64`, at most `2048`; a page costs an Office start and an export, so keeping one is worth it. `0` holds nothing between hovers, but a page is still rendered for the hover that asks for it.
- `pdf_cache_mb` — the same for the pages PDF previews were rendered as: default `32`, at most `2048`; a page is kept as the pixels it was drawn as, so the same file at two preview sizes is held as two pages.
- `text_cache_mb` — the same for the frames text previews were painted as: default `0`, at most `2048`; the text itself is already cached, and a frame is per file, per box and per scroll position — never one a selection was painted into.
- `decode_budget_gb` — the most memory one hover may decode or read for (a picture, the measurement of an SVG document and what a `.svgz` inflates to, a font's tables and what a `.woff2` inflates to, an animation's frames, the page Office exported, a theme file): default `1`, smallest `0.25`, largest `64`. It caps the file rather than what is kept, and a file past it shows no preview instead of the app asking for memory the allocator may refuse — which is why no value means "no limit".
- `office_engine_idle` — how long a family's Office engine is kept after that family's last page: seconds, or `indefinitely` for one kept as long as the app runs; default `600`. `0` lets it go as soon as it has drawn a page, so every document pays its own Office start.
- `office_extensions` — the Office-preview gate, under `[office]`, written without dots.
- `font_extensions` — the font-preview gate, under `[font]`, written without dots.
- A deleted `extensions=` line — or its whole section — comes back with the built-in entries; an `extensions=` line left empty stays empty.
- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled` — the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched at all; `true` by default.
- `follow_cursor` — `true` for Follow Cursor, `false` for Best Position.
- `avoid_mode` — `filename` (the default), `filename_column`, `details`, or `off`: what a preview is kept off, moving it and resizing it where there is no room beside that region, so the item it is about stays readable; applies to hovered and keyboard previews alike, at both positions. `filename` keeps it off the name alone, measured in the font the shell draws folder names in, at the scale of the display and at the larger of the sizes a view draws a name at (`Content` draws it 125% larger than a `Details` row), so the extension counts and only the name's own width is cleared — the rest of the column and the columns beside it are free to be covered; `filename_column` keeps it off the whole column the name sits in, as the view reports it; `details` keeps it off everything a row draws. A file written with the old `avoid_filename` key reads as `details` for `true` and `off` for `false`.
- `preview_scale` — percentage or `fit`.
- `svg_scale` — how much of the screen an SVG is drawn over: a percentage or `fit`, read against the screen rather than the size the document asks for, so `50` (the default) is half of it, `fit` all of it, and `100` or more read as `fit`. The engine's window is the size that comes out of this and its page fills it, so the setting is what the document is drawn at.
- `pdf_scale` / `office_scale` — the same for a PDF page and for a page Office rendered, both `fit` (the default) as well as a percentage; a workbook's fallback bitmap follows its share of its own size and is never enlarged.
- `font_scale` — how much of the screen a font specimen is drawn over: a percentage or `fit`, read against the screen, so `50` (the default) is half of it, `fit` all of it, and `100` or more read as `fit`. The specimen is sized from the window that comes out of this, so the setting is the size the type is drawn at. A font has no size of its own to be a percentage of, which is why the share is of the display rather than of the file.
- `font_background` — what a font specimen is drawn over: `black` (the default), `white`, `checkerboard`, or `transparent`. Text is drawn light on black, dark on the light backdrops, and light with a soft shadow where there is no backdrop at all to be read against.

## Build from Source

Requirements: Windows 11, Rust 1.88+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

```bash
cargo build            # debug
cargo build --release  # release
```

The release binary is written to `target/release/rust-hover-preview.exe`. A release build ends a running copy of the app first, since Windows will not let the linker replace a binary that is open; debug builds are left alone.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item, GDI paints the preview into a topmost layered window, text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives' own tables of contents, Office documents are drawn from the page Office renders in the background, the WebView2 runtime draws SVG documents and font specimens, and video is played by FFmpeg where it is installed and by Windows' own media engine where it is not.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview works fully offline — no telemetry, analytics, ads, update checks, or accounts. It reads only the item you hover or focus in Explorer, locally and only for enabled preview types. Cloud-only placeholders are skipped; password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use your local FFmpeg if you have installed it, Microsoft Office, Windows' own media engine, and the Windows PDF engine. Caches are in-memory and bounded by `config.ini`. See [PRIVACY.md](PRIVACY.md) for full details.

## License

MIT. See [LICENSE](LICENSE).

RAR archives are read with RARLAB's UnRAR sources, compiled into the binary by the `unrar` crate. UnRAR source code may be used in any software to handle RAR archives without limitations and free of charge, but it may not be used to develop a RAR-compatible archiver or to recreate the RAR compression algorithm, which is proprietary. See [RARLAB's licence](https://www.rarlab.com/license.htm) for the full terms.

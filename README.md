# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.88+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar. Hover a file in File Explorer — or move with the arrow keys — and a preview appears beside it. It works with mouse and keyboard, and you configure it from the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer.
- Images, including animated GIF, APNG, and WebP.
- HEIC, AVIF, and JPEG XL through Windows codec extensions where installed.
- SVG vectors drawn sharply at preview size.
- Font specimens: font name plus sample lines covering its own character map.
- Videos through FFmpeg if installed, otherwise through Windows’ own media engine.
- PDF first pages via the built-in Windows PDF engine.
- Text and code with syntax highlighting, rendered Markdown, and themes.
- Archives as a file tree with sizes — read without unpacking.
- Office documents drawn from a background Office render, so previews appear quickly after the first hover.
- Scaling from 25% to 400%, or fit-to-screen. Separate scaling for SVG, PDF, Office pages, and fonts.
- Previews appear beside the cursor or focused item and are kept on screen.
- Tray menu and hand-editable `config.ini`.
- DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.

## Supported Formats

You can add or remove formats in `config.ini`. Unsupported formats show no preview, except text, which the app will try to force-read.

### Images

`jpg`, `jpeg`, `png`, `apng`, `gif`, `bmp`, `ico`, `tiff`, `webp`, `tga`, `hdr`, `exr`, `qoi`, `heic`, `heif`, `avif`, `jxl`, `dds`, and more.

Animated GIF, APNG, and WebP files play. HEIC, HEIF, AVIF, JPEG XL, and still WebP usually need Windows codec extensions — except WebP, which the app can also decode on its own. HDR and EXR images are tone-mapped for screen preview. DDS textures preview in every block format they are written in — BC1 through BC7, plus unsigned BC6H — and in the uncompressed formats, showing the first face and first level of a cubemap or mip chain; the signed BC6H variant shows no preview.

### Vectors

`svg`, `svgz` — drawn by the WebView2 runtime Windows 11 ships with, so they stay sharp when enlarged. Animated SVG documents play too. If WebView2 is missing, SVG previews do not appear.

### Fonts

`ttf`, `otf`, `ttc`, `woff`, `woff2` — drawn by WebView2. The preview shows a font specimen: the name the font calls itself and sample lines its character map covers, from Latin through Arabic, Hebrew, Thai, and Devanagari.

### Videos

`mp4`, `webm`, `mkv`, `avi`, `mov`, `wmv`, `flv`, `m4v`, `ts`, `m2ts`, `mts`, `mpg`, `mpeg`, `vob`, `3gp`, `ogv`, `rmvb`, `asf`, `divx`, `f4v`, `mxf`, `dv`.

With FFmpeg installed, many more containers and codecs work. Without it, videos use the media engine Windows already has, plus any codec extensions you installed. The tray’s **Codecs** menu shows what is available.

### PDF

`pdf` — the first page is rendered by the Windows PDF engine. Password-protected and damaged files are skipped.

### Text and Code

`txt`, `md`, `rtf`, `nfo`, `json`, `toml`, `yaml`, `xml`, `ini`, `csv`, `log`, `sql`, `py`, `js`, `ts`, `rs`, `go`, `c`, `h`, `cpp`, `cs`, `java`, `kt`, `swift`, `php`, `rb`, `lua`, `sh`, `ps1`, `bat`, `html`, `css`, and more.

Extensionless files such as `LICENSE`, `Makefile`, `Dockerfile`, and `.gitignore` are also supported. Markdown can be rendered or shown as source. Full mode adds scrolling, selection, and copy.

### Archives

`zip`, `zipx`, `jar`, `apk`, `xpi`, `cbz`, `rar`, `7z`, `tar`, `tgz`, and `tar.gz`.

Archive previews show a file tree with sizes, read directly from the archive’s table of contents — nothing is unpacked.

### Office Documents

`doc`, `docm`, `docx`, `dot`, `dotm`, `dotx`, `xls`, `xlsb`, `xlsm`, `xlsx`, `xlt`, `xltm`, `xltx`, `ppt`, `pptm`, `pptx`, `pps`, `ppsm`, `ppsx`, `pot`, `potm`, `potx`.

Office previews require **Microsoft Office** installed. Excel needs a print queue to export a page — **Microsoft Print to PDF** is enough, and the Print Spooler service must be enabled. Without one, Excel falls back to the sheet’s top-left corner.

### Themes

Text, code, and archive listings use Atom One Light by default, One Dark Pro, or any `.tmTheme` file placed in:

```text
%APPDATA%\rust-hover-preview\theme
```

## Installation

Each release provides two options:

- `rust-hover-preview_<version>_x64-setup.exe` — NSIS installer. Installs to `%LOCALAPPDATA%\rust-hover-preview` with an optional startup entry.
- `rust-hover-preview.exe` — portable standalone binary. Run it from any folder.

Steps:

1. Open [Releases](../../releases).
2. Download your preferred asset.
3. Run the installer, or place the portable binary wherever you like.
4. Launch Rust Hover Preview.

No Rust toolchain is needed. If upgrading from an earlier version, the installer cleans up the old `%LOCALAPPDATA%\Rust Hover Preview` folder automatically.

## Optional: Enable Video Preview with FFmpeg

Videos preview without FFmpeg through the media engine Windows ships with. FFmpeg adds far more formats and codecs. Install it if you want formats Windows does not decode. `ffplay` and `ffprobe` need to be in your `PATH`.

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

## Optional: Enable More Video Codecs (Windows Codecs) — They are not needed while FFmpeg is installed.

The media engine decodes H.264, MPEG-4, and WMV out of the box. Each codec below is a separate free extension from the Microsoft Store:

| Codec                                  | Needs                                                                    |
| -------------------------------------- | ------------------------------------------------------------------------ |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)  |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)   |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)    |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4) |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)   |

All are free. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray’s **Codecs** menu is opened — no restart, nothing to configure.

## Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

`heic`, `heif`, `avif`, `jxl`, and still `webp` are decoded by a codec Windows provides rather than one shipped with the app. Each needs its extension installed once from the Microsoft Store:

| Format         | Needs                                                                                                                                            |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |

All are free. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error, and a multi-image file — a HEIC burst, an animated AVIF, an animated JPEG XL — shows its first frame. A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys or Tab — to preview the focused item.
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
  - **Avoid** — Don’t Avoid, Avoid Filename (default), Avoid Filename Column, or Avoid Details. This keeps a preview off the item it is about.
  - **Scaling** — Fit to Screen or 25%–400%.
  - **SVG Scaling** — Fit to Screen, or 75%, 50% (default), 25%, 10% of the display.
  - **PDF Scaling** — Fit to Screen (default), or the same shares of the display.
  - **Office Scaling** — the same for a page Office rendered; a workbook’s fallback bitmap is never enlarged.
  - **Font Scaling** — the same shares for a font specimen, 50% by default.
  - **Font Face** — which face of a `.ttc` collection is drawn: First Face (default) down to Tenth Face. The heading says which face came out, such as `(2 of 4)`.
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard.
  - **SVG Background** — the same backdrops for documents.
  - **Font Background** — the same backdrops for a font specimen.
  - **DDS Background** — the same backdrops for `.dds` textures, whose alpha channel is as often a mask or an unused channel as it is transparency.
- **Volume** — Max, High, Medium, Low, Very Low, Mute: 100% down to 0%.
- **Performance**
  - **Confirm File Type** — validate file content against the extension.
  - **Office Engine TTL** — how long a family’s Office app is kept warm: Indefinitely, 1 hour, 30 minutes, 10 minutes (default), 5 minutes, 1 minute, 0 seconds.
  - **SVG Engine TTL** — how long the browser that draws SVG documents is kept warm; greyed out where WebView2 is missing.
  - **Cache** — memory held between hovers, 2 GB down to 0 MB: Image, Text, PDF, Office caches.
  - **Decode Budget** — 16 GB down to 512 MB, 1 GB default: a file past it gets no preview.
- **Codecs** — what this machine has: Video, Images, and Engines. Missing ones are greyed out.
- **Run at Startup** — add or remove the Windows startup entry.
- **Config.ini** — open the configuration file; the item is named for the running version.
- **Exit** — close the app.

## Configuration

Settings are stored at:

```text
%APPDATA%\rust-hover-preview\config.ini
```

The file is watched, and changes apply without a restart.

Example, trimmed:

```ini
[settings]
run_at_startup=true
hover_delay_ms=0
same_file_rehover_delay_ms=750
spinner_delay_ms=250
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
hdr_tone_map=reinhard
hdr_exposure=0
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
text_preview_full_mode=false
text_font_scale=125

[image]
extensions=apng,avif,bmp,dds,exr,ff,gif,hdr,heic,heif,ico,jfif,jpe,jpeg,jpg,jxl,pam,pbm,pgm,png,pnm,ppm,qoi,svg,svgz,tga,tif,tiff,webp

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

Key settings, in plain terms:

- `theme` — `light`, `dark`, or `custom:<name>`.
- `markdown_mode` — `rendered` or `source`.
- `*_preview_enabled` — whether that kind of preview may show at all; the file lists stay as they are.
- `text_preview_full_mode` — `true` adds scrolling, selection, and copy.
- `text_font_scale` — a percentage from 1 to 1000, default `125`; archive listings follow it too.
- `extensions` / `names` — the text-preview gates. Extensions are written without dots. Names match extensionless files.
- `image_extensions`, `video_extensions`, `archive_extensions`, `office_extensions`, `font_extensions` — per-type preview gates, written without dots. An entry with a dot in it, like `tar.gz`, is matched against the end of the file name.
- `image_cache_mb` — memory for decoded image frames: default `32`, max `2048`; `0` holds nothing.
- `office_cache_mb` — memory for Office-rendered pages: default `64`, max `2048`; `0` holds nothing between hovers but still renders for the current hover.
- `pdf_cache_mb` — memory for PDF pages as pixels: default `32`, max `2048`; the same file at two preview sizes is held as two pages.
- `text_cache_mb` — memory for text frames: default `0`, max `2048`; the text itself is already cached.
- `decode_budget_gb` — the most memory one hover may decode or read for: default `1`, smallest `0.25`, largest `64`. A file past it shows no preview.
- `hdr_tone_map` — how HDR/EXR light values become screen values: `reinhard` (default), `aces` (filmic), `srgb` (clips), or `off` (bare clamp). Pictures already in screen values, like PNG or JPEG, are never affected.
- `hdr_exposure` — how many stops those pictures are shifted before that curve: default `0`, clamped to `-10`–`10`.
- `spinner_delay_ms` — how long a hover’s load may run before the waiting spinner is put up, in milliseconds: default `250`, and `0` puts it up with the load. One delay answers every kind of preview — a decode, a page Office is rendering, a browser that has to start — and there is no tray entry for it.
- `office_engine_idle` — seconds an Office engine is kept after its last page, or `indefinitely`; default `600`. `0` lets it go as soon as it has drawn a page.
- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled` — the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched at all; `true` by default.
- `follow_cursor` — `true` for Follow Cursor, `false` for Best Position.
- `avoid_mode` — `filename` (default), `filename_column`, `details`, or `off`: what a preview is kept off.
- `preview_scale` — percentage or `fit`.
- `svg_scale` — percentage or `fit`, read against the screen. `50` is default, `fit` is all of it, and `100` or more reads as `fit`.
- `pdf_scale` / `office_scale` — the same for a PDF page and an Office-rendered page, both `fit` by default as well as a percentage. A workbook’s fallback bitmap follows its own size and is never enlarged.
- `font_scale` — percentage or `fit`, read against the screen. `50` is default, `fit` is all of it, and `100` or more reads as `fit`. A font has no size of its own, so the share is of the display.
- `ttc_face` — which face of a `.ttc` collection is drawn: `1` is the first face, and the highest setting is `10`. The heading says which face came out.
- `font_background` — `black` (default), `white`, `checkerboard`, or `transparent`.
- A deleted `extensions=` line or whole section comes back with built-in entries. An `extensions=` line left empty stays empty.

## Build from Source

Requirements: Windows 11, Rust 1.88+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

```bash
cargo build            # debug
cargo build --release  # release
```

The release binary is written to `target/release/rust-hover-preview.exe`. A release build ends a running copy first, since Windows will not let the linker replace an open binary. Debug builds are left alone.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full system overview. In short: Windows accessibility APIs and Shell COM identify the hovered or focused Explorer item. GDI paints the preview into a topmost layered window. Text and code are highlighted with TextMate-style themes, Markdown is rendered, archive contents are listed from the archives’ own tables of contents, Office documents are drawn from a page Office renders in the background, WebView2 draws SVG documents and font specimens, and video is played by FFmpeg where installed and by Windows’ own media engine where it is not.

## TODO

See [TODO.md](TODO.md) for planned work, known bugs, and other issues.

## Privacy

Rust Hover Preview works fully offline — no telemetry, analytics, ads, update checks, or accounts. It reads only the item you hover or focus in Explorer, locally and only for enabled preview types. Cloud-only placeholders are skipped; password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use your local FFmpeg if installed, Microsoft Office, Windows’ own media engine, and the Windows PDF engine. Caches are in-memory and bounded by `config.ini`. See [PRIVACY.md](PRIVACY.md) for full details.

## License

MIT. See [LICENSE](LICENSE).

RAR archives are read with RARLAB’s UnRAR sources, compiled into the binary by the `unrar` crate. UnRAR source code may be used in any software to handle RAR archives without limitations and free of charge, but it may not be used to develop a RAR-compatible archiver or to recreate the RAR compression algorithm, which is proprietary. See [RARLAB’s licence](https://www.rarlab.com/license.htm) for the full terms.

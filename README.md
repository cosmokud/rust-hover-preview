# Rust Hover Preview

![Rust](https://img.shields.io/badge/Rust-1.98.1+-orange?logo=rust)
![Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)

A Windows 11 tray app inspired by QTTabBar. Hover a file in File Explorer — or move with the arrow keys — and a preview appears beside it. It works with mouse and keyboard, and you configure it from the tray icon or a simple `config.ini`.

[Showcase.webm](https://github.com/user-attachments/assets/33ee1f35-d399-4226-8847-5bd50f867ebb)

## Highlights

- Mouse-hover and keyboard-navigation previews in Explorer.
- Images, including animated GIF, APNG, and WebP.
- What a file really holds decides its preview: a video named as a document is played, a document named as a picture is drawn, and a file whose content no reader here answers for is left alone. A format whose bytes carry no signature — WordPerfect `wpd`, Lotus `wk4`, a RoQ video — is routed to the engine that reads it by extension, and a name that means two formats is settled by content: a `.pdb` is handed to LibreOffice only when it is a Palm OS ebook, never when it is a compiler's program database.
- Design documents — Photoshop, Illustrator, Krita, OpenRaster, and more — previewed from the picture their own format saves of the whole document.
- HEIC, AVIF, and JPEG XL through Windows codec extensions where installed.
- Vector drawings — SVG, Windows metafiles, and Illustrator `.eps` — drawn sharply at preview size.
- Font specimens: font name plus sample lines covering its own character map.
- Videos through FFmpeg if installed, otherwise through Windows’ own media engine.
- PDF first pages via the built-in Windows PDF engine.
- Text and code with syntax highlighting, rendered Markdown, and themes.
- Archives as a file tree with sizes — read without unpacking.
- Office documents drawn from a background Office render, so previews appear quickly after the first hover.
- Scaling from 25% to 400%, or fit-to-screen. Separate scaling for images and videos, and for vector drawings, PDF, Office pages, fonts, and design documents.
- Previews appear beside the cursor or focused item and are kept on screen.
- Tray menu and hand-editable `config.ini`.
- DPI aware, single-instance, sleep/resume resilient, and light on idle CPU.

## Supported Formats

You can add or remove formats in `config.ini`. Unsupported formats show no preview, except text, which the app will try to force-read.

Almost everything is previewed by **Windows 11 and this app alone** — a codec Windows ships, WebView2, the drawing layer that plays metafiles, the Windows PDF engine, or a reader written into the app. Three things are not, and they are the app’s largest optional dependencies:

### Runs on Windows 11 alone

| Kind                                       | Extensions                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Pictures                                   | `jpg` `jpeg` `jfif` `jpe` `png` `apng` `gif` `bmp` `ico` `tif` `tiff` `tga` `hdr` `exr` `ff` `qoi` `pnm` `pam` `pbm` `pgm` `ppm` `dds` `webp`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Pictures through a Windows codec extension | `heic` `heif` `avif` `jxl` — the extension packages are listed below; without one, no preview                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Videos Windows 11 plays itself             | `mp4` `m4v` `mov` `qt` `avi` `wmv` `asf` `dvr-ms` `mkv` `webm` `ts` `m2ts` `mts` `m2t` `mpg` `mpeg` `mpe` `m2v` `m1v` `vob` `3gp` `3g2` `3gpp`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Drawings                                   | `svg` `svgz` (WebView2) · `wmf` `emf` (the drawing layer) · `eps` `epsi` (the preview the file carries)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Design documents                           | `psd` `psb` `ai` `kra` `ora` `procreate` `sketch` `fig` `xd`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Fonts                                      | `ttf` `otf` `ttc` `woff` `woff2` (WebView2)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| PDFs                                       | `pdf` (the Windows PDF engine)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Text and code                              | `adb` `adoc` `ads` `asciidoc` `asm` `asp` `aspx` `astro` `awk` `bash` `bat` `bib` `bzl` `c` `cc` `cfg` `cg` `cjs` `clj` `cljc` `cljs` `cmake` `cmd` `comp` `conf` `cpp` `cs` `csh` `cshtml` `css` `csv` `csx` `cts` `cxx` `d` `dart` `diff` `diz` `edn` `ejs` `el` `elm` `env` `erb` `erl` `ex` `exs` `f` `f03` `f77` `f90` `f95` `fish` `for` `frag` `fs` `fsi` `fsx` `ftn` `fx` `geom` `glsl` `go` `gql` `gradle` `graphql` `groovy` `h` `haml` `hbs` `hcl` `hh` `hlsl` `hpp` `hrl` `hs` `htm` `html` `hxx` `inc` `ini` `ipynb` `java` `jl` `js` `json` `json5` `jsonc` `jsonl` `jsp` `jsx` `ksh` `kt` `kts` `latex` `less` `lhs` `liquid` `lisp` `ll` `lock` `log` `lsp` `lua` `m` `mak` `man` `markdown` `md` `mdown` `metal` `mjs` `mk` `mkd` `ml` `mli` `mm` `mts` `mustache` `nasm` `nfo` `nim` `ninja` `nix` `njk` `org` `pas` `patch` `php` `phtml` `pl` `plist` `pm` `properties` `proto` `ps1` `psd1` `psm1` `py` `pyi` `pyw` `r` `rake` `rb` `rkt` `rmd` `rs` `rst` `rtf` `s` `sass` `scala` `scm` `scss` `sh` `slim` `sol` `sql` `srt` `ss` `styl` `sv` `svelte` `svh` `swift` `tcl` `tex` `text` `tf` `tfvars` `toml` `ts` `tsv` `tsx` `twig` `txt` `v` `vbs` `vert` `vhd` `vhdl` `vtt` `vue` `wat` `wgsl` `xhtml` `xml` `xsd` `xsl` `xslt` `yaml` `yml` `zig` `zsh` — and these extensionless names: `authors` `.babelrc` `brewfile` `caddyfile` `changelog` `changes` `.clang-format` `.clang-tidy` `cmakelists.txt` `code_of_conduct` `containerfile` `contributing` `contributors` `copying` `copyright` `dockerfile` `.dockerignore` `.editorconfig` `.env` `.env.example` `.env.local` `.eslintignore` `.eslintrc` `gemfile` `.gitattributes` `.gitconfig` `.gitignore` `.gitkeep` `.gitmodules` `gnumakefile` `.golangci.yml` `history` `.htaccess` `install` `jenkinsfile` `justfile` `licence` `license` `.mailmap` `makefile` `makefile.am` `makefile.in` `notice` `.npmignore` `.prettierignore` `.prettierrc` `procfile` `rakefile` `readme` `.rustfmt.toml` `security` `.stylelintrc` `unlicense` `vagrantfile` |
| Archives                                   | `7z` `apk` `cbz` `jar` `rar` `tar` `tar.gz` `tgz` `xpi` `zip` `zipx`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |

### Notes on the table

Animated GIF, APNG and WebP play. `hdr` and `exr` are tone-mapped for preview. `dds` previews BC1–BC7, both BC6H variants and uncompressed textures — including packed HDR and depth — at the first face and the nearest-size mip.

### Needs FFmpeg, or the media engine Windows has

Videos the table does not already cover need FFmpeg, and these are all of them: `264` `265` `266` `apv` `av1` `avc` `avs` `avs2` `avs3` `bik` `bk2` `c93` `cavs` `cdg` `cdxl` `cin` `cpk` `dav` `dif` `divx` `drc` `dv` `evc` `f4v` `flm` `flv` `gxf` `h261` `h263` `h264` `h265` `h266` `h26l` `hevc` `ifv` `imx` `ismv` `ivf` `ivr` `kux` `m2p` `mj2` `mjpeg` `mjpg` `mk3d` `moflex` `mpv` `mve` `mvi` `mxf` `mxg` `nsv` `nut` `obu` `ogm` `ogv` `pmp` `psp` `rcv` `rm` `rmvb` `roq` `rsd` `smk` `str` `swf` `thp` `tod` `tp` `tr` `ty` `ty+` `usm` `vc1` `vc2` `viv` `vro` `vvc` `vw` `wtv` `xl` `xmv` `y4m` `yop`.

The containers and codecs the table above lists are the ones Windows plays on its own, and they play better with FFmpeg installed — more codecs inside the same container, and seeking rather than a still frame. The tray's **Codecs** menu shows what this machine can play.

### Needs Microsoft Office, or LibreOffice

Office documents — `doc` `docm` `docx` `dot` `dotm` `dotx` `xls` `xlsb` `xlsm` `xlsx` `xlt` `xltm` `xltx` `ppt` `pptm` `pptx` `pps` `ppsm` `ppsx` `pot` `potm` `potx` — are drawn in the background by an installed Office so previews appear quickly after the first hover. Excel needs a print queue to export a page — **Microsoft Print to PDF** is enough, and the Print Spooler service must be enabled; without one, Excel falls back to the sheet’s top-left corner. Where no Office is installed, an installed LibreOffice draws them instead. To have LibreOffice draw them all — with Microsoft Office installed as well — use **Engine → Select Engine → Office** and pick **LibreOffice**.

### Needs LibreOffice

The documents this app has no reader of its own for, drawn by an installed LibreOffice and sharp at any size, all of them: `123` `602` `abw` `cdr` `cgm` `cmx` `cwk` `dbf` `dif` `dxf` `fodg` `fodp` `fodt` `gnm` `gnumeric` `hwp` `key` `lwp` `mcw` `met` `mw` `numbers` `odb` `odc` `odf` `odg` `odm` `odp` `ods` `odt` `oth` `otg` `otm` `otp` `ots` `ott` `pages` `pcd` `pct` `pcx` `pdb` `pm6` `pmd` `psw` `pub` `ras` `sda` `sdc` `sdd` `sdw` `slk` `stc` `std` `sti` `stw` `svm` `sxd` `sxg` `sxi` `sxm` `sxw` `vdx` `vsd` `vsdm` `vsdx` `vstx` `wb2` `wk1` `wk3` `wk4` `wks` `wpg` `wq1` `wq2` `wpd` `wps` `wri` `xlw` `zabw` `zmf`.

Every name in it is one that a filter of the engine declares as something it imports, checked against the engine's own filter list rather than against the formats the engine is said to support. A name that no filter declares is a launch that answers nothing, which is why names an earlier version listed — `epub`, `qxp`, PageMaker before 6, the Visio stencils and templates, and a Flash file above all — are not in it, and a name added to it by hand is asked about from the next read.

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

### Optional: Enable Video Preview with FFmpeg

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

### Optional: Enable More Video Codecs (Windows Codecs) — They are not needed while FFmpeg is installed.

The media engine decodes H.264, MPEG-4, and WMV out of the box. Each codec below is a separate free extension from the Microsoft Store:

| Codec                                  | Needs                                                                    |
| -------------------------------------- | ------------------------------------------------------------------------ |
| HEVC (H.265)                           | [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ)  |
| VP9                                    | [VP9 Video Extensions](https://apps.microsoft.com/detail/9N4D0MSMP0PT)   |
| AV1                                    | [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)    |
| MPEG-1 and MPEG-2                      | [MPEG-2 Video Extension](https://apps.microsoft.com/detail/9N95Q1ZZPMH4) |
| Theora, Vorbis and Opus in an Ogg file | [Web Media Extensions](https://apps.microsoft.com/detail/9N5TDP8VCMHS)   |

All are free. Windows 11 usually has HEVC, VP9, and AV1 already. Installing one takes effect the next time the tray’s **Codecs** menu is opened — no restart, nothing to configure.

### Optional: Enable HEIC, AVIF, JPEG XL and WebP Preview (Windows Codecs)

`heic`, `heif`, `avif`, `jxl`, and still `webp` are decoded by a codec Windows provides rather than one shipped with the app. Each needs its extension installed once from the Microsoft Store:

| Format         | Needs                                                                                                                                            |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `heic`, `heif` | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [HEVC Video Extensions](https://apps.microsoft.com/detail/9N4WGH0Z6VHQ) |
| `avif`         | [HEIF Image Extension](https://apps.microsoft.com/detail/9PMMSR1CGPWG) + [AV1 Video Extension](https://apps.microsoft.com/detail/9MVZQVXJBQ9V)   |
| `jxl`          | [JPEG XL Image Extension](https://apps.microsoft.com/detail/9MZPRTH5C0TB), or the **JXL support** optional feature on Windows 11 24H2            |
| `webp`         | [WebP Image Extension](https://apps.microsoft.com/detail/9PG2DK419DRG) — optional: the app decodes WebP without it                               |

All are free. Windows 11 often has HEIF, AV1, and WebP already. Where one is missing, hovering such a file shows no preview rather than an error, and a multi-image file — a HEIC burst, an animated AVIF, an animated JPEG XL — shows its first frame. A `.webp` is the exception: it needs none of them, because the app carries its own libwebp decoder, so WebP previews even on Windows 10.

### Optional: Enable CorelDRAW and Other Documents (LibreOffice)

`cdr` and everything else in the `[libre]` list is drawn by **LibreOffice**, which this app drives where it finds it — nothing is bundled with the app and there is nothing to configure. Install it once:

```text
winget install -e --id TheDocumentFoundation.LibreOffice
```

If `winget` reports an error, the package sources are usually why: run `winget source reset --force`, and if that is refused as well, open **Terminal as administrator** (right-click the Start button → _Terminal (Admin)_) and run the same command from there. Then hover a `.cdr`: the first hover converts that document and takes a moment, and every hover after it is instant, because the converted page is kept beside `config.ini`. The engine itself is kept too, for ten minutes by default — **Engine → LibreOffice TTL** — so the next document does not pay for a start again.

## Usage

1. Start the app — a tray icon appears.
2. Hover media files in Explorer to preview them.
3. Or navigate with the keyboard — arrow keys or Tab — to preview the focused item.
4. Right-click the tray icon to configure behavior.

## System Tray Menu

The item a setting starts at carries `(Default)` after its name, so a menu says both what is set now — the check or radio mark — and what a setting nobody has touched would be.

- **Enable Preview** — turn previews on or off.
- **Preview Types** — Images, Videos, Text, PDF, Archives, Office, Vector, Fonts, Design: gate a kind without touching its file list.
- **Text Preview**
  - **Full Mode** — adds scrolling, selection, and copy; off by default.
  - **Theme** — Atom One Light, One Dark Pro, or any `.tmTheme` in the theme folder.
  - **Font Size** — 400% at the top down to 70% at the bottom.
  - **Markdown** — Rendered or Source.
- **Timing**
  - **Trigger Key (Alt)** — the key is named in the item itself.
    - **Enable Trigger Key** — whether the key is watched at all.
    - **Hold to Disable Preview** / **Hold to Enable Preview** — what holding it does.
  - **Delay** — how long the pointer rests on a file before its preview opens: 0 ms at the top down to 1000 ms at the bottom. 0 ms by default.
  - **Rehover Delay** — the same steps, before the same file can preview again; 200 ms by default.
  - **Settling Delay** — the same steps, how long the pointer must have been still before it will preview what it is on: 0 ms, the default, is a new file previewed while the hand is still moving to it. The keyboard's own previews are not gated by it.
- **Placement**
  - **Position** — Follow Cursor or Best Position; Best Position by default.
  - **Avoid** — Avoid Nothing, Avoid Filename (default), Avoid Filename Column, or Avoid Details. This keeps a preview off the item it is about.
- **Scaling**
  - **Image Scaling** — Fit to Screen or 25%–400%, of the image's own size.
  - **Video Scaling** — the same shares for a video, 100% (default).
  - **Animated Scaling** — the same shares for an animated GIF, WebP, or PNG; a still GIF or PNG keeps Image Scaling. 100% (default).
  - **Vector Scaling** — Fit to Screen (default), or 75%, 50%, 25%, 10% of the display.
  - **PDF Scaling** — Fit to Screen (default), or the same shares of the display.
  - **Office Scaling** — the same for a page Office rendered; a workbook’s fallback bitmap is never enlarged.
  - **Font Scaling** — the same shares for a font specimen, 50% by default.
  - **Design Scaling** — the same shares of the display for a design document (Photoshop, Illustrator, Krita, OpenRaster, Procreate); Fit to Screen by default.
  - **Libre Scaling** — the same shares of the display for a document LibreOffice draws (CorelDRAW and the older word processors, spreadsheets and presentations); Fit to Screen by default.
- **Background**
  - **Image Background** — Transparent, Black, White, or Checkerboard; Checkerboard by default.
  - **Vector Background** — the same backdrops for a drawing: an SVG document, or a metafile; Checkerboard by default.
  - **Font Background** — the same backdrops for a font specimen; White by default.
  - **DDS Background** — Black or White for `.dds` textures, whose alpha channel is as often a mask or an unused channel as it is transparency; White by default. The two backdrops that show what stands behind a texture are not offered for one.
  - **Design Background** — the same backdrops as a picture's, for a design document; Checkerboard by default.
- **Volume** — Max, High, Medium, Low, Very Low, Mute: 100% down to 0%.
- **Performance**
  - **Confirm File Type** — check a file's content against its name before it is previewed: a file whose bytes are another kind is previewed as that kind, and one whose content is a format this app has no reader for shows nothing. On by default.
  - **Cache** — memory held between hovers, 2 GB down to 0 MB: Image, Text, PDF, Office caches.
  - **Decode Budget** — 16 GB down to 512 MB, 1 GB default: a file past it gets no preview.
- **Engine**
  - **Select Engine → Office** — which engine an Office document’s page is asked of: **Microsoft Office** (default) draws it with the application that owns the format and keeps LibreOffice as the fallback for a family this machine has no application for, while **LibreOffice** draws every Office document whether Microsoft Office is installed or not. The LibreOffice row is greyed out where it is not installed.
  - **Microsoft Office TTL** — how long a family’s Office app is kept warm: Indefinitely, 1 hour, 30 minutes, 10 minutes (default), 5 minutes, 1 minute, 0 seconds.
  - **LibreOffice TTL** — the same for the engine that draws CorelDRAW and the documents beside it: how long it is kept after the last document it converted. A kept engine converts the next document without starting up again — measured on one document, 1.2 s cold against 0.2 s — and while it is kept it is a LibreOffice with a small document of this app's own open, which is a few hundred megabytes. `0 seconds` keeps none, which is an engine per document. Greyed out where LibreOffice is not installed.
  - **WebView2 TTL** — how long the browser that draws SVG documents is kept warm; greyed out where WebView2 is missing.
- **Codecs** — what this machine has: Videos, Images, and Engines. Missing ones are greyed out.
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
; General
preview_enabled=true
run_at_startup=true

; Preview Types
archive_preview_enabled=true
design_preview_enabled=true
font_preview_enabled=true
image_preview_enabled=true
office_preview_enabled=true
pdf_preview_enabled=true
text_preview_enabled=true
vector_preview_enabled=true
video_preview_enabled=true

; Text Preview
markdown_mode=rendered
text_font_scale=125
text_preview_full_mode=false
theme=light

; Timing
hover_delay_ms=0
same_file_rehover_delay_ms=200
settling_delay_ms=0
trigger_key=alt
trigger_key_enabled=true
trigger_key_mode=disable

; Placement
avoid_mode=filename
follow_cursor=false

; Scaling
animated_scale=100
design_scale=fit
font_scale=50
office_scale=fit
pdf_scale=fit
preview_scale=100
vector_scale=fit
video_scale=100

; Background
dds_background=white
design_background=checkerboard
font_background=white
image_background=checkerboard
vector_background=checkerboard

; Volume
video_volume=0

; Performance
confirm_file_type=true
decode_budget_gb=1
image_cache_mb=32
office_cache_mb=64
pdf_cache_mb=32
text_cache_mb=0

; Engine
libreoffice_idle=600
office_engine=microsoft_office
office_engine_idle=600

; Advanced
hdr_exposure=0
hdr_tone_map=reinhard
spinner_delay_ms=250
```

Key settings, in plain terms:

- `theme` — `light`, `dark`, or `custom:<name>`.
- `markdown_mode` — `rendered` or `source`.
- `*_preview_enabled` — whether that kind of preview may show at all; the file lists stay as they are.
- `text_preview_full_mode` — `true` adds scrolling, selection, and copy.
- `text_font_scale` — a percentage from 1 to 1000, default `125`; archive listings follow it too.
- `extensions` / `names` — the text-preview gates. Extensions are written without dots. Names match extensionless files.
- `image_extensions`, `video_extensions`, `archive_extensions`, `office_extensions`, `font_extensions`, `design_extensions`, `vector_extensions` — per-type preview gates, written without dots. An entry with a dot in it, like `tar.gz`, is matched against the end of the file name.
- `image_cache_mb` — memory for decoded image frames: default `32`, max `2048`; `0` holds nothing.
- `office_cache_mb` — memory for Office-rendered pages: default `64`, max `2048`; `0` holds nothing between hovers but still renders for the current hover.
- `pdf_cache_mb` — memory for PDF pages as pixels: default `32`, max `2048`; the same file at two preview sizes is held as two pages.
- `text_cache_mb` — memory for text frames: default `0`, max `2048`; the text itself is already cached.
- `decode_budget_gb` — the most memory one hover may decode or read for: default `1`, smallest `0.25`, largest `64`. A file past it shows no preview.
- `hdr_tone_map` — how HDR/EXR light values become screen values: `reinhard` (default), `aces` (filmic), `srgb` (clips), or `off` (bare clamp). Pictures already in screen values, like PNG or JPEG, are never affected.
- `hdr_exposure` — how many stops those pictures are shifted before that curve: default `0`, clamped to `-10`–`10`.
- `spinner_delay_ms` — how long a hover’s load may run before the waiting spinner is put up, in milliseconds: default `250`, and `0` puts it up with the load. One delay answers every kind of preview — a decode, a page Office is rendering, a browser that has to start — and there is no tray entry for it.
- `office_engine` — which engine draws an Office document’s page: `microsoft_office` (default) asks the application that owns the format and falls back to LibreOffice for a family this machine has no application for, while `libreoffice` asks the render engine for every Office document whether Microsoft Office is installed or not. With no LibreOffice installed the second falls back to the first, and the tray row is greyed out.
- `libreoffice_idle` — seconds the LibreOffice engine is kept after the last page it drew, or `indefinitely`; default `600`. `0` keeps no engine at all, which is a launch per document; while one is kept it is a LibreOffice with a small document of this app's own open, and it is ended by the app when the time is up.
- `office_engine_idle` — seconds an Office engine is kept after its last page, or `indefinitely`; default `600`. `0` lets it go as soon as it has drawn a page.
- `trigger_key` / `trigger_key_mode` / `trigger_key_enabled` — the key (`alt`, `ctrl`, `shift`, `win`), what it does (`disable` or `enable`), and whether it is watched at all; `true` by default.
- `follow_cursor` — `true` for Follow Cursor, `false` for Best Position.
- `avoid_mode` — `filename` (default), `filename_column`, `details`, or `off`: what a preview is kept off.
- `preview_scale` — percentage or `fit`, read against the picture's own size.
- `video_scale` — the same for a video, `100` by default.
- `animated_scale` — the same for an animated GIF, WebP, or PNG, `100` by default; a still GIF or PNG follows `preview_scale`.
- `vector_scale` — percentage or `fit`, read against the screen. `fit` is default, is all of the room, and `100` or more reads as it. It covers every drawing the Vector kind holds. The name it used to be written under (`svg_scale`) is not read: a line like that is removed the next time the file is written, and the setting goes back to its default.
- `libre_scale` — the same for a document LibreOffice draws, `fit` by default: the engine hands back a page, so the share is of the display the way a PDF page’s is.
- `libre_cache_mb` — megabytes of converted pages kept on disk, `32` by default; `0` keeps nothing between hovers. A page whose budget gave it up is converted again the next time it is hovered.
- `libre_preview_enabled` — whether documents LibreOffice can draw are previewed at all; `true` by default.
- `pdf_scale` / `office_scale` — the same for a PDF page and an Office-rendered page, both `fit` by default as well as a percentage. A workbook’s fallback bitmap follows its own size and is never enlarged.
- `font_scale` — percentage or `fit`, read against the screen. `50` is default, `fit` is all of it, and `100` or more reads as `fit`. A font has no size of its own, so the share is of the display.
- `design_scale` — the same for a design document, `fit` by default. A design preview is the picture the file keeps of the whole document, so the share is of the screen the way a page’s is rather than of the document’s own size.
- `ttc_face` — which face of a `.ttc` collection is drawn: `1` is the first face, and the highest setting is `10`. The heading says which face came out.
- `image_background` — `checkerboard` (default), `black`, `white`, or `transparent`: what a picture is drawn over, and with it a PDF page, a painted text frame and a page Office rendered. The one name every kind used to share (`transparent_background`) is not read: lines like that are removed the next time the file is written.
- `font_background` — `white` (default), `black`, `checkerboard`, or `transparent`.
- `dds_background` — `white` (default) or `black`, and only those two: a texture's alpha channel is as often a mask, a height or a roughness as it is transparency, so the backdrops that show what stands behind a preview are not offered for one, and a file that names one of them is read as `white`.
- `vector_background` — `checkerboard` (default), `white`, `black`, or `transparent`: what an SVG document's page, or a metafile drawing, is drawn over. The name it used to be written under (`svg_background`) is not read: a line like that is removed the next time the file is written.
- `design_background` — `checkerboard` (default), `white`, `black`, or `transparent`.
- A deleted `extensions=` line or whole section comes back with built-in entries. An `extensions=` line left empty stays empty.

## Build from Source

Requirements: Windows 11, Rust 1.98.1+, Visual Studio Build Tools (MSVC / C++), and the Windows SDK.

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

Rust Hover Preview works fully offline — no telemetry, analytics, ads, accounts, or crash reporting. It reads only the item you hover or focus in Explorer, locally and only for enabled preview types. Cloud-only placeholders are skipped on purpose; password-protected files are never bypassed. Settings and themes live under `%APPDATA%\rust-hover-preview`; optional previews use your local FFmpeg if installed, Microsoft Office, Windows’ own media engine, and the Windows PDF engine. Caches are in-memory and bounded by `config.ini`. The only network request is an update check, which runs only when you open the tray menu and at most once an hour. See `PRIVACY.md` for full details.

## License

MIT. See [LICENSE](LICENSE).

The app is built out of other people’s code as much as its own — the Windows bindings, the image decoders, the syntax highlighter, the archive readers, the browser bindings — and each of those carries its own licence, with the notices MIT and BSD ask to be reproduced. [THIRD-PARTY.md](THIRD-PARTY.md) lists every dependency grouped by licence, with the copyright holders beside it, and the full texts are in [`LICENSES/`](LICENSES). Both are generated from the dependency tree rather than kept by hand; `generate-attribution.ps1` refreshes them.

RAR archives are read with RARLAB’s UnRAR sources, compiled into the binary by the `unrar` crate. UnRAR source code may be used in any software to handle RAR archives without limitations and free of charge, but it may not be used to develop a RAR-compatible archiver or to recreate the RAR compression algorithm, which is proprietary. See [RARLAB’s licence](https://www.rarlab.com/license.htm) for the full terms.
